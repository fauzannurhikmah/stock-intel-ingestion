
from flask import Blueprint, request, jsonify

from app.validators import (
    validate_emiten_request,
    validate_fundamental_request,
    validate_shares_data_request,
    validate_technical_request,
)
from app.services.fundamental_service import fetch_and_build_fundamental
from app.services.emiten_service import scrape_emiten_detail, scrape_emiten_list
from app.services.shares_service import fetch_and_build_shares_data
from app.serializers.emiten_serializer import serialize_emiten_detail, serialize_emiten_list
from utils.technical import fetch_technical_analysis


bp = Blueprint("main", __name__, url_prefix="/api")


@bp.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


@bp.route("/fundamental", methods=["POST"])
def get_fundamental():
    body = request.get_json(silent=True) or {}
    symbol, year, quarter, errors = validate_fundamental_request(body)
    if errors:
        return jsonify({"status": "error", "errors": errors}), 400

    try:
        response = fetch_and_build_fundamental(symbol, year, quarter)
        return jsonify(response)
    except Exception:
        return jsonify({
            "status": "error",
            "message": "Failed to retrieve data from IDX. Please try again later."
        }), 502


@bp.route("/technical-analysis", methods=["POST"])
def get_technical_analysis():
    body = request.get_json(silent=True) or {}

    emiten, errors = validate_technical_request(body)
    if errors:
        return jsonify({
            "status": "error",
            "errors": errors,
        }), 400

    try:
        result = fetch_technical_analysis(emiten)
        return jsonify({
            "status": "ok",
            "input": {"emiten": emiten},
            "technical_analysis": result,
        })
    except ValueError as exc:
        return jsonify({
            "status": "error",
            "message": str(exc),
        }), 400
    except Exception:
        return jsonify({
            "status": "error",
            "message": "Failed to fetch technical analysis. Please try again later.",
        }), 502


@bp.route("/emiten", methods=["GET"])
def get_emiten():
    query = request.args.to_dict(flat=True)
    symbol, page, page_size, sort_type, sort_direction, errors = validate_emiten_request(query)
    if errors:
        return jsonify({"status": "error", "errors": errors}), 400

    try:
        if symbol:
            result = scrape_emiten_detail(symbol)
            return jsonify(serialize_emiten_detail(result))

        result = scrape_emiten_list(
            page=page,
            page_size=page_size,
            sort_type=sort_type,
            sort_direction=sort_direction,
        )
        return jsonify(serialize_emiten_list(result))
    except ValueError as exc:
        return jsonify({"status": "error", "message": str(exc)}), 404
    except Exception:
        return jsonify({
            "status": "error",
            "message": "Failed to fetch emiten data. Please try again later.",
        }), 502


@bp.route("/shares-data", methods=["GET"])
def get_shares_data():
    query = request.args.to_dict(flat=True)
    symbol, errors = validate_shares_data_request(query)
    if errors:
        return jsonify({"status": "error", "errors": errors}), 400

    try:
        result = fetch_and_build_shares_data(symbol)
        return jsonify(result)
    except ValueError as exc:
        return jsonify({"status": "error", "message": str(exc)}), 404
    except Exception:
        return jsonify({
            "status": "error",
            "message": "Failed to fetch shares data from IDX. Please try again later.",
        }), 502
